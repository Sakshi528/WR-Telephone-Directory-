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
        ("Database",             "PostgreSQL (via psycopg[binary] v3)"),
        ("ORM",                  "Flask-SQLAlchemy 3.1.1"),
        ("Authentication",       "Flask-Login 0.6.3"),
        ("Excel Export",         "openpyxl 3.1.5"),
        ("Word Export",          "python-docx 1.1.2"),
        ("Frontend Framework",   "Bootstrap 5.3 (vendored locally under static/ -- no CDN dependency, "
                                  "so the app works on LAN machines without internet access)"),
        ("Icons",                "Bootstrap Icons 1.11 (vendored locally under static/)"),
        ("Colour Scheme",        "Navy Blue (#1A3A6B) on white / Bootstrap neutrals"),
        ("Session Management",   "Flask server-side sessions with 20-minute idle timeout"),
    ]
    for row in stack:
        add_table_row(tbl, list(row))

    doc.add_paragraph()

    # ── 3. Database Structure ─────────────────────────────────────────────────────
    add_heading(doc, "3.  Database Structure", level=1)
    doc.add_paragraph(
        "The application uses a PostgreSQL database with 19 tables in total. The schema is "
        "defined through SQLAlchemy models (models/*.py) plus numbered, hand-written SQL "
        "migration files in migrations/ (applied with psql -- no Alembic). The seven original "
        "tables (users, employees, organizations, departments, directory_numbers, "
        "update_requests, emergency_contacts) pre-date the migration system; every other table "
        "was added by a numbered migration (002 to 019). Below is every table and its purpose."
    )

    tbl2 = doc.add_table(rows=1, cols=3)
    tbl2.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl2.style     = "Table Grid"
    add_table_row(tbl2, ["#  Table", "Key Columns", "Purpose"], header=True)

    tables_info = [
        ("1  users",
         "id, username, email (unique), password (hashed), role, status, created_at",
         "Admin login accounts only. Passwords stored as Werkzeug pbkdf2 hashes. role = 'admin' "
         "grants access. There is no admin screen to create/edit users -- accounts are created "
         "with create_admin.py."),
        ("2  employees",
         "id, employee_name, designation, organization_id (FK), department_id (FK), location, "
         "region, office_phone, residence_phone, mobile_phone, email, is_utility_head, is_kmp, "
         "status, status_changed_at",
         "Every named contact. is_kmp separates region-wide Key Managerial Personnel contacts "
         "from WRLDC-internal staff. status is one of ACTIVE, INACTIVE, TRANSFERRED, RETIRED, "
         "DEPUTATION, RESIGNED, CONTRACT_ENDED, DECEASED. designation (free text) drives the "
         "automatic Utility Head detection."),
        ("3  departments",
         "id, department_name (unique)",
         "Lookup table for internal departments (e.g. Operations, IT, HR)."),
        ("4  organizations",
         "id, organization_name (unique), region, state, address, parent_id (self-FK), "
         "category_id (FK), utility_head_excluded",
         "Every power utility, generator, load despatch centre and government body, with a "
         "parent / sub-organisation hierarchy. state drives the Type -> State -> Organisation "
         "cascading filters. The Subcategory value is stored in the region column. utility_head_excluded lets an admin mark an organisation as having "
         "no Utility Head, overriding automatic resolution."),
        ("5  organization_categories",
         "id, category_name (unique), description, is_state_based",
         "Controlled 'Organisation Type' list (Transmission Utility, Generation Company, "
         "Thermal, Hydel, Nuclear, State SLDC, DISCOM, CPSU, Others, etc.). is_state_based "
         "decides whether a State step appears in the cascading filters. Admin-managed from "
         "/admin/categories."),
        ("6  organization_subcategories",
         "id, category_id (FK, ON DELETE CASCADE), subcategory_name",
         "Reusable Subcategory suggestion list per category, used for autocomplete on the "
         "Add/Edit Organisation form. Never a required reference (migration 018)."),
        ("7  service_types",
         "id, service_type_name (unique), description",
         "Civil-service cadre lookup (IAS, IPS, IFS, State Civil Service, Other) attached to "
         "Administrative Heads for filtering/reporting only."),
        ("8  directory_numbers",
         "id, name, phone_number, email, organization, organization_id (FK), category, "
         "switch_yard, control_room, ip_address",
         "Control room numbers, switchyard numbers, ALDC lines and other numbers not tied to "
         "a person. Linked to organizations via organization_id so each number groups under "
         "the right organisation card."),
        ("9  emergency_contacts",
         "id, employee_id (FK), contact_name, relation, phone",
         "Next-of-kin / emergency contact of each non-KMP (Grid India) employee."),
        ("10 administrative_heads",
         "id, organization_id (FK), employee_id (FK, nullable), role_category, role_title, "
         "service_type_id (FK), effective_from, status, remarks, standalone name / phones / "
         "email / office_address, created_by, created_at",
         "Current holder of a senior post (Administrative Head, or KMP such as Company "
         "Secretary) at an organisation. May be linked to an Employee or be 'standalone' "
         "(e.g. a government-appointed Principal Secretary with no employee record). Only one "
         "Administrative Head per organisation + role title is allowed."),
        ("11 administrative_head_assistants",
         "id, administrative_head_id (FK, cascade), employee_id (FK), designation, name, "
         "office_phone, mobile, email, is_primary",
         "PA / PS attached to an Administrative Head. At most one primary assistant per head."),
        ("12 administrative_head_history",
         "snapshot of the head row + effective_to, reason, replacement_employee_id",
         "Immutable archive written when an Administrative Head's role is ended."),
        ("13 employee_status_history",
         "id, employee_id (FK, cascade), old/new status, old/new organisation, old/new "
         "designation, reason, changed_by, changed_at",
         "Change log per employee (status, organisation, designation) -- feeds the Status "
         "History timeline."),
        ("14 update_requests",
         "id, employee_id / directory_number_id / administrative_head_id (FK), request_type, "
         "field_name, new_value, requested_by, department, contact_number, reason, status, "
         "reviewed_by, reviewed_at, admin_comment, ip_address, browser, operating_system",
         "Anonymous public correction requests for an Employee, Directory Number or "
         "Administrative Head. Admin approves or rejects."),
        ("15 audit_logs",
         "id, module, record_type, record_id, action, field_name, old_value, new_value, "
         "changed_by, reason, ip_address, browser, operating_system, session_id, created_at",
         "Append-only change log for admin actions. Never updated or deleted by the app."),
        ("16 import_batches",
         "id, workbook_name, imported_by, import_date, mode (DRY_RUN / APPLIED), "
         "inserted / updated / skipped counts, status, summary",
         "One row per data-import script run (shown on the Import History screen)."),
        ("17 directory_versions",
         "id, version_number (unique), version_name, month, year, generated_on, generated_by, "
         "employee_count, organization_count, pdf / excel filename and path, status, remarks",
         "One row per generated official monthly directory snapshot (PDF + Excel). A "
         "regeneration always creates a new suffixed version; older ones are never modified."),
        ("18 email_groups  (+ group_members)",
         "id, name (unique), description, list_type (STATIC / DYNAMIC), created_by; "
         "group_members: group_id, employee_id",
         "Saved custom email distribution groups. DYNAMIC groups are computed live from "
         "filters; STATIC groups keep an explicit member list in group_members."),
        ("19 email_group_filters",
         "id, group_id (FK), filter_type, organization_id, category_id, role_value, "
         "status_value, employee_id",
         "One filter rule for a DYNAMIC group: ORGANIZATION, ORGANIZATION_CATEGORY, ROLE "
         "(Utility Head / Administrative Head / KMP), STATUS, plus INCLUDE_EMPLOYEE / "
         "EXCLUDE_EMPLOYEE manual overrides. Same type = OR, different types = AND."),
    ]
    for row in tables_info:
        add_table_row(tbl2, list(row))

    doc.add_paragraph()
    add_heading(doc, "3.1  Delete Behaviour (Foreign Keys)", level=2)
    for bold, desc in [
        ("Organisation deleted",   "Blocked by the admin screen while it still has employees, "
                                   "directory numbers, child organisations or administrative heads. "
                                   "Its own category link is SET NULL."),
        ("Category deleted",       "Blocked while any organisation uses it; its subcategory "
                                   "suggestions are deleted with it (cascade)."),
        ("Employee deleted",       "Emergency contacts, status history and email-group overrides "
                                   "linked to that employee are removed with it (cascade)."),
        ("Administrative Head ended", "Archived into administrative_head_history, then the live row "
                                      "and its PA/PS assistants are deleted."),
    ]:
        add_bullet(doc, desc, bold)

    add_heading(doc, "3.2  Migration History", level=2)
    tblm = doc.add_table(rows=1, cols=2)
    tblm.alignment = WD_TABLE_ALIGNMENT.LEFT
    tblm.style     = "Table Grid"
    add_table_row(tblm, ["Migration", "What it did"], header=True)
    for row in [
        ("add_parent_id_and_org_fk", "organizations.parent_id (hierarchy) and directory_numbers.organization_id."),
        ("002 enterprise_features", "Employee status + history, administrative heads, email groups, audit logs, update-request review columns."),
        ("003 email_distribution_lists", "organization_categories (10 seed types), organizations.category_id, email group list_type and filters."),
        ("004 contact_indexes", "Performance indexes on employee / directory-number organisation links."),
        ("005 administrative_heads_enhancement", "service_types, service type on heads, PA/PS assistants."),
        ("006 administrative_heads_standalone", "Heads can exist without an Employee record."),
        ("007 archive_module", "Extra employee statuses, IP/browser/OS audit metadata, import_batches."),
        ("008 directory_versions", "Monthly PDF + Excel snapshot table."),
        ("009-012 organisation types", "State-based flag, organizations.state, new types (CTU, STU, State SLDC, CPSU, DISCOM, Thermal, Hydel, Nuclear, Generation Company)."),
        ("013 update_request_administrative_head", "Public correction requests can target Administrative Heads."),
        ("014 directory_number_contact_fields", "switch_yard / control_room flags and ip_address on directory numbers."),
        ("015 remove_ipp_category", "Removed the unused IPP category."),
        ("016 utility_head_no_head_flag", "organizations.utility_head_excluded."),
        ("017 email_group_member_overrides", "INCLUDE_EMPLOYEE / EXCLUDE_EMPLOYEE manual overrides for custom groups."),
        ("018 organization_subcategories", "Subcategory suggestion table."),
        ("019 remove_re_generators_category", "Removed the empty RE Generators category (its 74 organisations were merged into Generation Company on 2026-08-27)."),
    ]:
        add_table_row(tblm, list(row))

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
        "scoring (Chairman/CMD = highest, Manager = lowest) via "
        "utils/designation_rank.py, with two admin overrides on top: an admin can "
        "manually pin a specific employee as an organisation's head regardless of "
        "designation, or explicitly mark an organisation as having no Utility Head at "
        "all (instead of always falling back to the same employee when there is no "
        "other candidate)."
    )
    items = [
        ("Columns",           "Sr No., Name, Designation, Organisation, Office Phone, "
                              "Mobile, Email, Address."),
        ("Organisation address", "The registered address of each organisation is shown in "
                                 "the Address column."),
        ("Export – Excel",    "Download the complete utility heads list as .xlsx."),
        ("Export – Word",     "Download the complete utility heads list as .docx."),
        ("Admin controls",    "Edit, Remove from Utility Heads, and Delete buttons appear "
                              "for admin users. Removing the head of a single-employee (or "
                              "otherwise uncontested) organisation excludes that "
                              "organisation from having a Utility Head, rather than "
                              "silently re-selecting the same employee."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    add_heading(doc, "4.4.1  How Utility Heads Are Detected", level=3)
    doc.add_paragraph(
        "Utility Heads are never flagged by hand. utils/designation_rank.py ranks every "
        "employee's free-text designation against an ordered list of about 30 seniority tiers "
        "(Chairman, Chairman & Managing Director, Managing Director, Director, CEO, Executive "
        "Director, Chief General Manager, General Manager, Chief Engineer, ... down to Operator). "
        "Each title and its abbreviation (e.g. GM / CGM / ED / CMD) map to the same tier, and the "
        "longest matching phrase wins so 'Additional Chief Engineer' is not mistaken for "
        "'Chief Engineer'. For each organisation the most senior ACTIVE employee becomes its "
        "Utility Head, recalculated on every page load. To change who appears, edit the employee's "
        "designation (or mark the organisation as 'no Utility Head' in the admin screen)."
    )

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

    # 4.8 Email Lists
    add_heading(doc, "4.9  Email Distribution Lists  (/email-lists)", level=2)
    doc.add_paragraph(
        "Lets anyone build a mailing list of the right people in a few clicks -- e.g. every "
        "Utility Head, every Thermal generator contact in a state, or an admin-defined custom "
        "group. Nothing here needs a login."
    )
    items = [
        ("Home page",        "/email-lists shows cards grouped in sections: Organisation Based "
                             "(one card per organisation category with contact counts; opens the "
                             "Address Book), Role Based (Utility Heads, Administrative Heads, KMP "
                             "etc.; opens a distribution-list page) and Custom Groups (admin-created "
                             "groups, each showing member and unique-email counts)."),
        ("Address Book",     "/email-lists/browse?category=... -- a fast grid of every contact in "
                             "a category. Filters: State (only for state-based categories), Contact "
                             "Type (e.g. only Utility Heads or only Administrative Heads) and a "
                             "live search box. Rows are fetched as JSON and filtered instantly in "
                             "the browser without reloading. Admins additionally see Add / Edit / "
                             "Delete on each row."),
        ("Distribution List","/email-lists/<group> -- a table of contacts for a role-based list "
                             "with an optional contact-type filter."),
        ("Custom Group page","/email-lists/custom/<id> -- the current members and emails of an "
                             "admin-created custom group, with the same copy / compose tools."),
        ("Row selection",    "Every row has a checkbox, plus Select All, so only the chosen people "
                             "are used by the copy and compose buttons."),
        ("Copy Selected Emails", "Copies only the ticked rows' email addresses to the clipboard, "
                             "duplicates removed, sorted alphabetically, separated by semicolons "
                             "(ready to paste into Outlook's To: field)."),
        ("Copy All / Name+Email / Full Details", "Copy every email, copy 'Name <email>' pairs, or "
                             "copy full contact details (name, designation, organisation, phones, "
                             "email)."),
        ("Compose Email",    "Opens the user's default mail app with all selected (or all listed) "
                             "addresses pre-filled in the To: field using a mailto: link. Very long "
                             "recipient lists cannot fit in a mailto: link, so the app then shows a "
                             "pop-up offering to copy the emails instead."),
        ("Export",           "Download the same list as Excel (.xlsx) or CSV; the file always "
                             "matches the filters currently on screen."),
        ("De-duplication",   "Someone who is both a Utility Head and an Administrative Head of the "
                             "same organisation appears only once."),
        ("Always live",      "Custom groups are recalculated from current data on every visit, so "
                             "a new joiner or a status change is reflected immediately."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 4.10 Administrative Heads
    add_heading(doc, "4.10  Administrative Heads  (/administrative-heads)", level=2)
    for bold, desc in [
        ("List",     "Senior post holders (Administrative Heads and KMP such as Company Secretary) "
                     "with the same Type -> State -> Organisation cascading filters, plus service "
                     "type, status and designation/keyword filters. 'Standalone' badge marks heads "
                     "who have no employee record."),
        ("Detail page", "/administrative-heads/<id> -- Official Contact, PA/PS contacts and Office "
                     "Address cards, and a Request Update link."),
        ("How they are decided", "Not automatic: an admin assigns each head. A head linked to an "
                     "employee always shows that employee's live name/phone/email; a standalone "
                     "head shows details stored on the head record itself."),
        ("Export",   "/export/administrative-heads downloads the current list."),
    ]:
        add_bullet(doc, desc, bold)

    # 4.11 Directory Versions
    add_heading(doc, "4.11  Directory Versions  (/directory-versions)", level=2)
    for bold, desc in [
        ("Purpose",  "Official month-by-month snapshots of the whole directory, each available as "
                     "a PDF and an Excel file."),
        ("Browse",   "Filter by month, year or keyword (paginated). Each row has Preview (PDF in "
                     "browser), Download PDF and Download Excel."),
        ("Contents", "The Excel has 7 sheets: Employees, Utility Heads, Administrative Heads, KMP, "
                     "Control Rooms, Switchyards, Emergency Contacts. Navy header styling is shared "
                     "with the PDF and web UI."),
    ]:
        add_bullet(doc, desc, bold)

    # 4.8 Universal Search
    add_heading(doc, "4.12  Universal Search  (/search)", level=2)
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
        ("Add / Edit",     "Organisation name, region, category, optional subcategory (stored in region) "
                           "(autocomplete from the category's suggestion list) and physical address."),
        ("Delete",         "Remove an organisation (cascades to employees)."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 5.4 Directory Numbers
    add_heading(doc, "5.4  Directory Numbers  (/admin/directory-numbers)", level=2)
    items = [
        ("Search",        "Search by name, phone, organisation, or category."),
        ("Add / Edit",    "Name, phone number, email, Organisation (dropdown of actual "
                          "organisations -- links via organization_id so it groups "
                          "correctly on the Telephone Directory page), category "
                          "(e.g. Control Room, Switchyard). Both the Add and Edit forms "
                          "now carry a Back button that returns to this list."),
        ("Delete",        "Remove a directory number entry."),
        ("Dashboard link","Back to dashboard button in the page header."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 5.7 Custom Email Groups
    add_heading(doc, "5.5  Custom Email Groups  (/admin/email-groups)", level=2)
    doc.add_paragraph(
        "Saved, filter-based distribution lists. A group is defined by Organisation "
        "Category / Organisation, Role (Utility Head / Administrative Head / KMP), and "
        "Status filters -- membership is never a stored list, it is resolved live from "
        "the current employee data every time the group is viewed or emailed."
    )
    items = [
        ("Add / Edit group",   "Choose Organisation Categories or specific Organisations, "
                               "Roles, and Statuses. Rows within one filter type are OR'd "
                               "together; different filter types are AND'd."),
        ("Members column",     "Two badges on the group list show the total matched "
                               "member count and the unique-email count; both are "
                               "clickable and open the group's Members page."),
        ("View members",       "The Members page lists every currently matched contact "
                               "(Organisation, Name, Designation, Email) with a live "
                               "search box, resolved fresh on every visit."),
        ("Manually add a member",    "From the Members page, an admin can add a specific "
                                     "employee to a group even if they don't match its "
                                     "filters -- this creates an explicit override that "
                                     "always applies, regardless of the filter criteria."),
        ("Manually remove a member", "Likewise, an admin can remove a specific employee "
                                     "from a group even if they DO match its filters -- "
                                     "the removal sticks until reversed, independent of "
                                     "future filter or data changes."),
        ("Public visibility", "Custom groups also appear to all users on /email-lists as \"Custom Groups\" cards, opening /email-lists/custom/<id>."),
        ("Delete group",       "Remove a saved custom group entirely."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 5.5 Emergency Contacts
    add_heading(doc, "5.6  Emergency Contacts  (/admin/emergency-contacts)", level=2)
    items = [
        ("Search",        "Search by employee name, contact name, or relation."),
        ("Add / Edit",    "Link to employee, contact name, relation, phone."),
        ("Delete",        "Remove a contact record."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 5.6 Update Requests
    add_heading(doc, "5.7  Update Requests  (/admin/requests)", level=2)
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

    # 5.8 Organization Categories
    add_heading(doc, "5.8  Organisation Categories  (/admin/categories)", level=2)
    doc.add_paragraph(
        "Lets an admin register a new Organisation Category (e.g. ahead of a new type of "
        "station being added) or manage a reusable Subcategory suggestion list per category, "
        "without needing a code change or database migration. Reached via the \"Add Category\" "
        "card on the Admin Dashboard."
    )
    items = [
        ("Add Category",       "Name, optional description, and a \"state-based\" flag "
                               "(controls whether a State step appears in that category's "
                               "cascading Type -> State -> Organisation filters)."),
        ("Add Subcategory",    "A free-text suggestion registered under a Category -- shows "
                               "up as an autocomplete suggestion on the Add/Edit Organisation "
                               "form's Subcategory field. Subcategories don't need to exist "
                               "before an organisation uses that value; this list is purely a "
                               "convenience, never a required reference."),
        ("Remove Category / Subcategory", "A Category can't be deleted while any organisation "
                                          "still uses it -- reassign them first. A Subcategory "
                                          "suggestion can always be removed freely (it's never "
                                          "itself referenced by an organisation)."),
        ("Sync from live data", "Reconciles the Subcategory suggestion list against what "
                                "organisations actually carry right now -- adds any missing "
                                "combination in live use, and removes suggestions nothing "
                                "matches any more (e.g. left over after a bulk "
                                "recategorisation elsewhere)."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 5.9 Administrative Heads
    add_heading(doc, "5.9  Administrative Heads  (/admin/administrative-heads)", level=2)
    for bold, desc in [
        ("Assign Head",  "Link an existing employee (autocomplete) OR enter a standalone person's "
                         "name and contact details. Choose category (Administrative Head or KMP), "
                         "role title, service type, effective-from date and status (Active, On "
                         "Leave, Vacant, Inactive)."),
        ("Edit",         "Update any field; a summary box shows what is currently displayed across "
                         "the app and warns when an employee link overrides standalone details."),
        ("End Role",     "Requires a reason and optional replacement; archives a full snapshot to "
                         "administrative_head_history and removes the live record and its PA/PS."),
        ("PA / PS",      "Add, edit and delete assistants for a head; 'Set as Primary' (one primary "
                         "per head)."),
    ]:
        add_bullet(doc, desc, bold)

    # 5.10 Directory Versions
    add_heading(doc, "5.10  Directory Versions  (admin actions)", level=2)
    for bold, desc in [
        ("Generate",   "Creates the PDF + Excel snapshot for a month/year (blocked if one already "
                       "exists; use Regenerate). A live version-number preview is shown."),
        ("Regenerate", "Always allowed; creates a new suffixed version (e.g. 2026.08.2) and never "
                       "changes the older one."),
        ("Remarks / Delete", "Edit remarks, or delete a version (removes the database row and its "
                       "files from disk)."),
    ]:
        add_bullet(doc, desc, bold)

    # 5.11 Employee status
    add_heading(doc, "5.11  Employee Status & History", level=2)
    for bold, desc in [
        ("Status change", "From the employee detail page an admin can set ACTIVE, INACTIVE, "
                          "TRANSFERRED, RETIRED, DEPUTATION, RESIGNED, CONTRACT_ENDED or DECEASED "
                          "with a reason. Non-active employees are greyed out with a status badge "
                          "and dropped from active lists, email lists and Utility Head detection."),
        ("Status History", "/admin/employees/<id>/status-history lists every change of status, "
                           "organisation or designation with old/new values, reason, who and when."),
        ("KMP / Utility Head toggles", "One-click toggles on the employee list; KMP changes are "
                           "audit-logged."),
        ("Approved requests", "An approved public 'status' correction request goes through the same "
                           "logic as an admin status change."),
    ]:
        add_bullet(doc, desc, bold)

    # 5.12 Import history / verification
    add_heading(doc, "5.12  Import History, Audit Log & Verification Workbook", level=2)
    for bold, desc in [
        ("Import History", "/admin/import-history -- read-only list (50 per page) of data-import "
                           "runs with mode (dry run / applied), counts and status. Only imports run "
                           "after this feature was added are recorded."),
        ("Audit log",      "Admin changes are written to audit_logs with old/new value, user, IP, "
                           "browser and OS. There is currently no screen to view it."),
        ("Verification Workbook", "/admin/export-verification-workbook -- Excel with the default "
                           "view of every public page (organisations, employees, telephone "
                           "directory, utility heads, administrative heads, control rooms, "
                           "switchyards) for pre-deployment checking."),
    ]:
        add_bullet(doc, desc, bold)

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
        ("Administrative Heads",       "Excel (.xlsx)",
         "Current filtered list of administrative heads."),
        ("Email lists / Address Book", "Excel (.xlsx) / CSV",
         "Contacts of a list or category exactly as filtered on screen."),
        ("Directory Version",          "PDF + Excel (7 sheets)",
         "Monthly official snapshot, stored on disk and downloadable."),
        ("Verification Workbook",      "Excel (.xlsx), admin only",
         "Default view of every public page for deployment checking."),
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
        ("/admin/email-groups",                "Admin",  "Manage custom email groups"),
        ("/admin/email-groups/add",            "Admin",  "Add custom email group"),
        ("/admin/email-groups/edit/<id>",      "Admin",  "Edit custom email group"),
        ("/admin/email-groups/delete/<id>",    "Admin",  "Delete custom email group"),
        ("/admin/email-groups/<id>/members",   "Admin",  "View a group's resolved members"),
        ("/admin/email-groups/<id>/members/add",    "Admin", "Manually add an employee to a group"),
        ("/admin/email-groups/<id>/members/remove", "Admin", "Manually remove an employee from a group"),
        ("/administrative-heads",             "Public", "Administrative Heads list"),
        ("/administrative-heads/<id>",         "Public", "Administrative Head detail"),
        ("/directory-versions",                "Public", "Monthly directory snapshots"),
        ("/directory-versions/<id>/preview | download/pdf | download/excel", "Public", "View / download a snapshot"),
        ("/email-lists/<slug>",                "Public", "Role-based distribution list"),
        ("/email-lists/<slug>/emails.json | export.xlsx | export.csv", "Public", "List emails (JSON) and exports"),
        ("/email-lists/browse",                "Public", "Address Book"),
        ("/email-lists/browse/data.json | export.xlsx | export.csv", "Public", "Address Book data and exports"),
        ("/export/administrative-heads",       "Public", "Export Administrative Heads"),
        ("/suggestions | /utility-heads/suggest", "Public", "Search autocomplete (JSON)"),
        ("/api/organization-types/<id>/states | /api/organizations", "Public", "Cascading dropdown data (JSON)"),
        ("/session-expired",                   "Public", "Idle-timeout page"),
        ("/admin/employees/<id>/status | status-history | toggle-kmp", "Admin", "Employee status change, history, KMP toggle"),
        ("/admin/employees/suggest",           "Admin",  "Employee picker suggestions"),
        ("/admin/administrative-heads/add | edit/<id> | end/<id>", "Admin", "Assign, edit, end an Administrative Head"),
        ("/admin/administrative-heads/<id>/assistants/add | edit | delete | set-primary", "Admin", "Manage PA/PS assistants"),
        ("/admin/directory-versions/generate | <id>/regenerate | <id>/remarks | <id>/delete", "Admin", "Manage directory versions"),
        ("/admin/import-history",              "Admin",  "Import history"),
        ("/admin/export-verification-workbook","Admin",  "Verification workbook"),
        ("/admin/categories",                  "Admin",  "Manage organisation categories & subcategories"),
        ("/admin/categories/add | edit/<id> | delete/<id>", "Admin", "Add, edit or delete a category"),
        ("/admin/categories/<id>/subcategories/add", "Admin", "Add a subcategory suggestion"),
        ("/admin/categories/subcategories/delete/<id>", "Admin", "Remove a subcategory suggestion"),
        ("/admin/categories/sync",             "Admin",  "Sync suggestions from live data"),
        ("/api/organization-categories/<id>/subcategories", "Admin", "Subcategory suggestions (JSON)"),
        ("/email-lists",                       "Public", "Email distribution lists home (incl. Custom Groups cards)"),
        ("/email-lists/custom/<id>",           "Public", "View an admin-created custom group's members & emails"),
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
        ("templates/manage_email_groups.html", "Admin – custom email group list, with clickable member/email-count badges."),
        ("templates/email_group_form.html",    "Admin – add/edit a custom email group's filters."),
        ("templates/email_group_members.html", "Admin – a group's live-resolved member list, with manual add/remove."),
        ("models/organization_subcategory.py", "OrganizationSubcategory model (suggestion list per category)."),
        ("templates/manage_categories.html",   "Admin – categories & subcategory suggestions."),
        ("templates/email_distribution_custom_group.html", "Public – a custom group's members and emails."),
        ("utils/designation_rank.py",          "Keyword-based designation ranking used to auto-resolve Utility Heads."),
        ("migrations/018_organization_subcategories.sql", "Creates organization_subcategories."),
        ("migrations/019_remove_re_generators_category.sql", "Removes the empty 'RE Generators' category "
                                                            "(its orgs were merged into Generation Company)."),
        ("models/email_group.py",              "EmailGroup model (static or dynamic/filter-based distribution list)."),
        ("models/email_group_filter.py",       "EmailGroupFilter model -- one filter criterion per group, plus the "
                                                "INCLUDE_EMPLOYEE/EXCLUDE_EMPLOYEE manual-override rows."),
        ("services/email_distribution_service.py", "Resolves a group's live membership from its filters and "
                                                    "manual overrides."),
        ("models/administrative_head*.py",     "AdministrativeHead, AdministrativeHeadAssistant, AdministrativeHeadHistory models."),
        ("models/service_type.py, audit_log.py, employee_status_history.py, import_batch.py, directory_version.py", "Remaining models (service types, audit log, status history, import batches, directory versions)."),
        ("services/head_service.py",           "Assign / update / end administrative head roles; search."),
        ("services/directory_version_service.py, excel_generator.py, pdf_generator.py", "Generate monthly PDF + Excel snapshots."),
        ("services/audit_service.py, org_stats.py, verification_workbook_service.py", "Audit logging, organisation statistics, verification workbook."),
        ("static/js/",                         "autocomplete.js, org_type_cascade.js, email_distribution.js, email_distribution_browse.js (copy / compose / select logic)."),
        ("utils/",                             "designation_rank.py (Utility Head detection), logger.py, user_agent.py, section_order.py."),
        ("scripts/",                           "Data validation, reconciliation, reclassification and audit scripts (see section 11)."),
        ("static/css/style.css",      "Custom styles: navy colour scheme, org card, badges."),
        ("create_admin.py",           "One-time script to create the admin user in the DB."),
        ("import_employee.py",        "Script to import employees from the Word document."),
        ("import_directory_numbers.py","Script to import control room numbers."),
        ("update_addresses_and_heads.py","Script to extract org addresses and mark utility heads."),
        ("fix_data_quality.py",       "Script to correct org names, spacing, and Unicode errors."),
        ("database/postgres_schema.sql","Base PostgreSQL DDL; later changes are in migrations/*.sql (up to 019)."),
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

    add_heading(doc, "11.  Organisation Classification & Data Import", level=1)
    doc.add_paragraph(
        "Organisations are classified with a single category (organization_categories) and an "
        "optional free-text subcategory, plus a state. The Type -> State (only for state-based "
        "types) -> Organisation cascading dropdown is used consistently on the Employee, "
        "Administrative Head, Utility Heads and Administrative Heads screens."
    )
    doc.add_paragraph(
        "Data originally came from the official Word directory (2026 Main_TD.docx). It was "
        "repaired and validated into an Excel workbook (scripts/repair_workbook.py, "
        "validate_workbook.py, compare_word_excel.py, reconcile_word_database.py, "
        "apply_corrections.py, normalize_formatting.py) and then loaded into PostgreSQL "
        "(import_from_excel_db.py for organisations, KMP contacts, control rooms and "
        "switchyards; import_employees_excel.py for WRLDC staff from uploads/employee list.xlsx, "
        "which is the only source for WRLDC employees). New organisations default to 'Others' "
        "and are then classified with scripts/classify_organizations.py and "
        "reclassify_organization_types.py. Ongoing audits use validate_database.py and the "
        "reports/ folder (e.g. category/subcategory bifurcation, Word-vs-app field diff)."
    )

    add_heading(doc, "12.  Known Limitations", level=1)
    for bold, desc in [
        ("User management", "No screen to create or manage admin accounts (use create_admin.py)."),
        ("Audit log",       "Recorded but there is no viewer screen."),
        ("Emergency Contacts page", "The public page exists but is not linked from the main menu."),
        ("Import rollback", "Import runs are logged but cannot be rolled back."),
        ("Public access",   "All read-only pages are open to anyone on the network; only edits are protected."),
    ]:
        add_bullet(doc, desc, bold)

    out = "Project_Summary_Telephone_Directory.docx"
    doc.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    build()
