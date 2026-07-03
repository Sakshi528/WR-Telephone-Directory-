import io
from datetime import datetime

from flask import (
    Blueprint, jsonify, render_template, request,
    redirect, url_for, flash, send_file
)
from flask_login import current_user
from sqlalchemy import or_

from models import db
from models.employee import Employee
from models.organization import Organization
from models.department import Department
from models.directory_number import DirectoryNumber
from models.emergency_contact import EmergencyContact
from models.update_request import UpdateRequest


user_bp = Blueprint("user", __name__)


# ─── SEARCH FILTERS ──────────────────────────────────────────────────────────

def employee_search_filter(keyword):
    filters = [
        Employee.employee_name.ilike(f"%{keyword}%"),
        Employee.designation.ilike(f"%{keyword}%"),
        Employee.region.ilike(f"%{keyword}%"),
        Organization.organization_name.ilike(f"%{keyword}%"),
        Organization.region.ilike(f"%{keyword}%"),
        Organization.address.ilike(f"%{keyword}%"),
        Department.department_name.ilike(f"%{keyword}%"),
        Employee.location.ilike(f"%{keyword}%"),
        Employee.office_phone.ilike(f"%{keyword}%"),
        Employee.mobile_phone.ilike(f"%{keyword}%"),
    ]
    # Only search email when the keyword looks like an email address,
    # otherwise "ntpc" matches every @ntpc.co.in employee across all companies.
    if "@" in keyword:
        filters.append(Employee.email.ilike(f"%{keyword}%"))
    return or_(*filters)


def directory_number_search_filter(keyword):
    from sqlalchemy import select as sa_select
    # match dir numbers whose org string belongs to an org in the searched region
    region_org_names = sa_select(Organization.organization_name).where(
        Organization.region.ilike(f"%{keyword}%")
    )
    filters = [
        DirectoryNumber.name.ilike(f"%{keyword}%"),
        DirectoryNumber.phone_number.ilike(f"%{keyword}%"),
        DirectoryNumber.category.ilike(f"%{keyword}%"),
        DirectoryNumber.organization.ilike(f"%{keyword}%"),
        DirectoryNumber.organization.in_(region_org_names),
    ]
    if "@" in keyword:
        filters.append(DirectoryNumber.email.ilike(f"%{keyword}%"))
    return or_(*filters)


def get_matched_organizations(keyword):
    return Organization.query.filter(
        or_(
            Organization.organization_name.ilike(f"%{keyword}%"),
            Organization.region.ilike(f"%{keyword}%"),
            Organization.address.ilike(f"%{keyword}%"),
        )
    ).all()


# ─── SUGGESTION HELPERS ──────────────────────────────────────────────────────

def add_suggestion(suggestions, seen, label, value, category):
    if not label or not value:
        return
    key = value.strip().lower()
    if key in seen:
        return
    seen.add(key)
    suggestions.append({"label": label, "value": value, "category": category})


def ranked_items(items, keyword, label_getter):
    keyword = keyword.lower()
    return sorted(
        items,
        key=lambda item: (
            not label_getter(item).lower().startswith(keyword),
            label_getter(item).lower(),
        ),
    )


# ─── HOME ────────────────────────────────────────────────────────────────────

@user_bp.route("/")
def home():
    # If already logged-in admin, redirect to dashboard
    if current_user.is_authenticated and current_user.role == "admin":
        return redirect(url_for("admin.dashboard"))

    if request.args.get("expired"):
        flash("Your session expired due to inactivity.", "warning")

    total_employees = Employee.query.count()
    total_organizations = Organization.query.count()

    # Last updated = most recently created employee record (proxy for data freshness)
    last_emp = Employee.query.order_by(Employee.id.desc()).first()
    last_updated = last_emp.id if last_emp else None   # placeholder; real date needs created_at

    return render_template(
        "index.html",
        total_employees=total_employees,
        total_organizations=total_organizations,
        last_updated=datetime.today().strftime("%B %Y"),
    )


@user_bp.route("/session-expired")
def session_expired():
    flash("Your session expired due to inactivity.", "warning")
    return redirect(url_for("user.home"))


# ─── EMPLOYEE DIRECTORY (WRLDC ONLY) ─────────────────────────────────────────

@user_bp.route("/directory")
def directory():
    keyword = request.args.get("keyword", "")

    query = Employee.query.outerjoin(Organization).outerjoin(Department)
    query = query.filter(Employee.is_kmp == False)  # noqa: E712

    if keyword:
        query = query.filter(employee_search_filter(keyword))

    employees = query.order_by(Employee.employee_name).all()

    # Build emergency-contact lookup: employee_id → first EmergencyContact
    emp_ids = [e.id for e in employees]
    ec_map = {}
    if emp_ids:
        contacts = EmergencyContact.query.filter(
            EmergencyContact.employee_id.in_(emp_ids)
        ).all()
        for c in contacts:
            if c.employee_id not in ec_map:
                ec_map[c.employee_id] = c

    return render_template(
        "directory.html",
        employees=employees,
        keyword=keyword,
        ec_map=ec_map,
    )


@user_bp.route("/employee/<int:id>")
def employee_details(id):
    employee = db.session.get(Employee, id) or _abort404()
    return render_template("employee_details.html", employee=employee)


# ─── UPDATE REQUESTS (ANONYMOUS) ─────────────────────────────────────────────

@user_bp.route("/employee/request-update/<int:id>", methods=["GET", "POST"])
def request_update(id):
    employee = db.session.get(Employee, id) or _abort404()
    if request.method == "POST":
        update_request = UpdateRequest(
            user_id=None,
            employee_id=employee.id,
            request_type="employee",
            requested_by=request.form.get("requested_by"),
            department=request.form.get("department"),
            contact_number=request.form.get("contact_number"),
            field_name=request.form.get("field_name"),
            new_value=request.form.get("new_value"),
            reason=request.form.get("reason"),
        )
        db.session.add(update_request)
        db.session.commit()
        flash("Update request submitted. Thank you!", "success")
        return redirect(url_for("user.employee_details", id=employee.id))
    return render_template("update_request.html", employee=employee)


@user_bp.route(
    "/directory-number/request-update/<int:id>", methods=["GET", "POST"]
)
def request_directory_update(id):
    number = db.session.get(DirectoryNumber, id) or _abort404()
    if request.method == "POST":
        update_request = UpdateRequest(
            user_id=None,
            directory_number_id=number.id,
            request_type="directory",
            requested_by=request.form.get("requested_by"),
            department=request.form.get("department"),
            contact_number=request.form.get("contact_number"),
            field_name=request.form.get("field_name"),
            new_value=request.form.get("new_value"),
            reason=request.form.get("reason"),
        )
        db.session.add(update_request)
        db.session.commit()
        flash("Directory correction request submitted. Thank you!", "success")
        return redirect(url_for("user.directory_number_details", id=number.id))
    return render_template("directory_update_request.html", number=number)


# ─── TELEPHONE DIRECTORY (ALL ORGS) ──────────────────────────────────────────

@user_bp.route("/telephone-directory")
def telephone_directory():
    keyword = request.args.get("keyword", "")
    employees = []
    control_rooms = []
    other_numbers = []
    matched_organizations = []

    if not keyword:
        # Build org address lookup
        all_orgs = Organization.query.all()
        org_address = {o.organization_name: (o.address or "") for o in all_orgs}

        all_employees = (
            Employee.query
            .outerjoin(Organization)
            .filter(Employee.is_kmp == True)  # noqa: E712
            .order_by(
                Organization.organization_name,
                Employee.is_utility_head.desc(),
                Employee.employee_name,
            )
            .all()
        )
        all_numbers = (
            DirectoryNumber.query
            .filter(DirectoryNumber.organization != None)
            .order_by(DirectoryNumber.organization)
            .all()
        )

        grouped = {}
        for emp in all_employees:
            org_name = (
                emp.organization.organization_name if emp.organization else "Unknown"
            )
            if org_name not in grouped:
                grouped[org_name] = {
                    "employees": [],
                    "control_rooms": [],
                    "switchyards": [],
                    "address": emp.organization.address if emp.organization else "",
                }
            grouped[org_name]["employees"].append(emp)

        default_other = []
        for dn in all_numbers:
            org_name = dn.organization or "Unknown"
            if org_name not in grouped:
                grouped[org_name] = {
                    "employees": [],
                    "control_rooms": [],
                    "switchyards": [],
                    "address": org_address.get(org_name, ""),
                }
            if dn.category and "switchyard" in dn.category.lower():
                grouped[org_name]["switchyards"].append(dn)
            elif dn.category and "control room" in dn.category.lower():
                grouped[org_name]["control_rooms"].append(dn)
            else:
                default_other.append(dn)

        return render_template(
            "telephone_directory.html",
            grouped=dict(sorted(grouped.items())),
            other_numbers=default_other,
            matched_organizations=[],
            keyword="",
            employees=all_employees,
            control_rooms=[],
            numbers=default_other,
        )

    matched_organizations = get_matched_organizations(keyword)
    matched_org_names = [o.organization_name for o in matched_organizations]
    matched_org_ids   = [o.id for o in matched_organizations]

    emp_query = (
        Employee.query.outerjoin(Organization).outerjoin(Department)
        .filter(Employee.is_kmp == True)  # noqa: E712
    )
    if matched_org_ids:
        employees = emp_query.filter(
            or_(
                Employee.organization_id.in_(matched_org_ids),
                employee_search_filter(keyword),
            )
        ).order_by(
            Organization.organization_name,
            Employee.is_utility_head.desc(),
            Employee.employee_name,
        ).limit(200).all()
    else:
        employees = emp_query.filter(
            employee_search_filter(keyword)
        ).order_by(
            Employee.is_utility_head.desc(),
            Employee.employee_name,
        ).limit(200).all()

    cr_filters = [
        DirectoryNumber.name.ilike(f"%{keyword}%"),
        DirectoryNumber.phone_number.ilike(f"%{keyword}%"),
    ]
    if "@" in keyword:
        cr_filters.append(DirectoryNumber.email.ilike(f"%{keyword}%"))
    for org_name in matched_org_names:
        if org_name:
            cr_filters.append(DirectoryNumber.organization.ilike(f"%{org_name}%"))

    all_matched_numbers = DirectoryNumber.query.filter(or_(*cr_filters)).all()
    control_rooms = [
        n for n in all_matched_numbers
        if n.category and "control room" in n.category.lower()
    ]
    switchyards = [
        n for n in all_matched_numbers
        if n.category and "switchyard" in n.category.lower()
    ]
    other_numbers = [
        n for n in all_matched_numbers
        if not (n.category and "control room" in n.category.lower())
        and not (n.category and "switchyard" in n.category.lower())
    ]

    # Build org address lookup
    org_address = {o.organization_name: (o.address or "") for o in matched_organizations}

    grouped = {}
    for emp in employees:
        org_name = (
            emp.organization.organization_name if emp.organization else "Unknown"
        )
        if org_name not in grouped:
            grouped[org_name] = {
                "employees": [],
                "control_rooms": [],
                "switchyards": [],
                "address": emp.organization.address if emp.organization else "",
            }
        grouped[org_name]["employees"].append(emp)

    for cr in control_rooms:
        org_name = cr.organization or "Unknown"
        if org_name not in grouped:
            grouped[org_name] = {
                "employees": [],
                "control_rooms": [],
                "switchyards": [],
                "address": org_address.get(org_name, ""),
            }
        grouped[org_name]["control_rooms"].append(cr)

    for sw in switchyards:
        org_name = sw.organization or "Unknown"
        if org_name not in grouped:
            grouped[org_name] = {
                "employees": [],
                "control_rooms": [],
                "switchyards": [],
                "address": org_address.get(org_name, ""),
            }
        grouped[org_name]["switchyards"].append(sw)

    # When a parent org matches the keyword, include ALL its sub-orgs (even empty),
    # so that e.g. "Black Start" shows every Black Start sub-org with its state.
    child_orgs = set()
    for org in matched_organizations:
        for child in Organization.query.filter_by(region=org.organization_name).all():
            child_orgs.add(child.organization_name)
            if child.organization_name not in grouped:
                grouped[child.organization_name] = {
                    "employees": [],
                    "control_rooms": [],
                    "switchyards": [],
                    "address": child.address or child.region or "",
                }

    # Show orgs that have content, or that are explicitly sub-orgs of a matched parent.
    grouped_sorted = dict(sorted(
        (k, v) for k, v in grouped.items()
        if v["employees"] or v["control_rooms"] or v["switchyards"] or k in child_orgs
    ))

    return render_template(
        "telephone_directory.html",
        grouped=grouped_sorted,
        other_numbers=other_numbers,
        matched_organizations=matched_organizations,
        keyword=keyword,
        employees=employees,
        control_rooms=control_rooms,
        numbers=other_numbers,
    )


@user_bp.route("/directory-number/<int:id>")
def directory_number_details(id):
    number = db.session.get(DirectoryNumber, id) or _abort404()
    return render_template("directory_number_details.html", number=number)


# ─── UNIVERSAL SEARCH ─────────────────────────────────────────────────────────

@user_bp.route("/search")
def universal_search():
    keyword = request.args.get("keyword", "")
    employees = []
    directory_numbers = []

    if keyword:
        employees = Employee.query.outerjoin(Organization).outerjoin(
            Department
        ).filter(employee_search_filter(keyword)).all()

        directory_numbers = DirectoryNumber.query.filter(
            directory_number_search_filter(keyword)
        ).all()

    return render_template(
        "universal_search.html",
        employees=employees,
        directory_numbers=directory_numbers,
        keyword=keyword,
    )


# ─── UTILITY HEADS ────────────────────────────────────────────────────────────

@user_bp.route("/utility-heads")
def utility_heads():
    keyword = request.args.get("keyword", "").strip()

    query = (
        Employee.query
        .outerjoin(Organization)
        .filter(Employee.is_utility_head == True)  # noqa: E712
    )
    if keyword:
        query = query.filter(
            Organization.organization_name.ilike(f"%{keyword}%")
        )
    heads = query.order_by(Organization.organization_name, Employee.employee_name).all()

    # Deduplicate by (lowercase name + org id) to remove accidental duplicate records
    seen = set()
    deduped = []
    for emp in heads:
        key = (emp.employee_name.strip().lower(), emp.organization_id)
        if key not in seen:
            seen.add(key)
            deduped.append(emp)

    return render_template("utility_heads.html", heads=deduped, keyword=keyword)


@user_bp.route("/utility-heads/suggest")
def utility_heads_suggest():
    q = request.args.get("q", "").strip()
    if len(q) < 1:
        return jsonify([])
    orgs = (
        Organization.query
        .join(Employee, Organization.id == Employee.organization_id)
        .filter(
            Employee.is_utility_head == True,  # noqa: E712
            Organization.organization_name.ilike(f"%{q}%"),
        )
        .with_entities(Organization.organization_name)
        .distinct()
        .order_by(Organization.organization_name)
        .limit(10)
        .all()
    )
    return jsonify([row[0] for row in orgs])


# ─── EMERGENCY CONTACTS ───────────────────────────────────────────────────────

@user_bp.route("/emergency-contacts")
def emergency_contacts():
    contacts = (
        EmergencyContact.query
        .join(Employee)
        .order_by(Employee.employee_name)
        .all()
    )
    return render_template("emergency_contacts.html", contacts=contacts)


# ─── AUTOCOMPLETE SUGGESTIONS ─────────────────────────────────────────────────

@user_bp.route("/suggestions")
def search_suggestions():
    keyword = request.args.get("q", "").strip()
    source  = request.args.get("source", "all")
    suggestions = []
    seen = set()

    if len(keyword) < 2:
        return jsonify(suggestions)

    if source in {"all", "employee"}:
        organizations = Organization.query.filter(
            or_(
                Organization.organization_name.ilike(f"%{keyword}%"),
                Organization.region.ilike(f"%{keyword}%"),
            )
        ).limit(20).all()

        for org in ranked_items(organizations, keyword, lambda o: o.organization_name)[:6]:
            add_suggestion(suggestions, seen,
                           org.organization_name, org.organization_name, "Organization")

        departments = Department.query.filter(
            Department.department_name.ilike(f"%{keyword}%")
        ).limit(12).all()

        for dept in ranked_items(departments, keyword, lambda d: d.department_name)[:4]:
            add_suggestion(suggestions, seen,
                           dept.department_name, dept.department_name, "Department")

        emps = Employee.query.outerjoin(Organization).outerjoin(
            Department
        ).filter(employee_search_filter(keyword)).limit(24).all()

        for emp in ranked_items(emps, keyword, lambda e: e.employee_name)[:8]:
            label = emp.employee_name
            if emp.designation:
                label = f"{label} - {emp.designation}"
            add_suggestion(suggestions, seen, label, emp.employee_name, "Employee")

    if source in {"all", "directory"}:
        numbers = DirectoryNumber.query.filter(
            directory_number_search_filter(keyword)
        ).limit(24).all()

        for number in ranked_items(numbers, keyword, lambda n: n.name)[:8]:
            label = number.name
            if number.phone_number:
                label = f"{label} - {number.phone_number}"
            add_suggestion(suggestions, seen, label, number.name,
                           number.category or "Directory")

    return jsonify(suggestions[:10])


# ─── EXPORT ───────────────────────────────────────────────────────────────────

@user_bp.route("/export/employees")
def export_employees():
    fmt      = request.args.get("format", "xlsx")
    keyword  = request.args.get("keyword", "")
    org_id   = request.args.get("organization", "")
    dept_id  = request.args.get("department", "")

    wrldc_ids = [
        o.id for o in Organization.query.filter(
            or_(
                Organization.organization_name.ilike("%WRLDC%"),
                Organization.organization_name.ilike("%Western Region Load Despatch%"),
                Organization.organization_name.ilike("%Western Region Load Dispatch%"),
            )
        ).all()
    ]

    query = Employee.query.outerjoin(Organization).outerjoin(Department)
    query = query.filter(Employee.organization_id.in_(wrldc_ids))
    if keyword:
        query = query.filter(employee_search_filter(keyword))
    if org_id:
        query = query.filter(Employee.organization_id == org_id)
    if dept_id:
        query = query.filter(Employee.department_id == dept_id)

    employees = query.order_by(Employee.employee_name).all()

    if fmt == "docx":
        return _export_employees_docx(employees)
    return _export_employees_xlsx(employees)


@user_bp.route("/export/telephone")
def export_telephone():
    fmt     = request.args.get("format", "xlsx")
    keyword = request.args.get("keyword", "")

    # Reuse same query logic as telephone_directory view
    if keyword:
        matched_organizations = get_matched_organizations(keyword)
        matched_org_ids = [o.id for o in matched_organizations]

        emp_query = (
            Employee.query.outerjoin(Organization).outerjoin(Department)
            .filter(Employee.is_kmp == True)  # noqa: E712
        )
        if matched_org_ids:
            employees = emp_query.filter(
                or_(
                    Employee.organization_id.in_(matched_org_ids),
                    employee_search_filter(keyword),
                )
            ).order_by(Organization.organization_name, Employee.employee_name).all()
        else:
            employees = emp_query.filter(
                employee_search_filter(keyword)
            ).order_by(Employee.employee_name).all()
    else:
        employees = (
            Employee.query.outerjoin(Organization)
            .filter(Employee.is_kmp == True)  # noqa: E712
            .order_by(Organization.organization_name, Employee.employee_name)
            .all()
        )

    if fmt == "docx":
        return _export_employees_docx(employees, title="Telephone Directory")
    return _export_employees_xlsx(employees, title="Telephone Directory")


@user_bp.route("/export/org-word/<path:org_name>")
def export_org_word(org_name):
    employees = (
        Employee.query
        .join(Organization)
        .filter(Organization.organization_name == org_name)
        .order_by(Employee.employee_name)
        .all()
    )
    control_rooms = (
        DirectoryNumber.query
        .filter(
            DirectoryNumber.organization.ilike(f"%{org_name}%"),
            DirectoryNumber.category == "Control Room",
        )
        .all()
    )
    org = Organization.query.filter_by(organization_name=org_name).first()
    return _export_org_docx(org_name, org.address if org else "", employees, control_rooms)


# ─── EXPORT HELPERS ───────────────────────────────────────────────────────────

def _export_employees_xlsx(employees, title="Employee Directory"):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title[:31]

    headers = ["Name", "Designation", "Organization", "Department",
               "Office Phone", "Mobile", "Email", "Utility Head"]
    header_font   = Font(bold=True, color="FFFFFF")
    header_fill   = PatternFill("solid", fgColor="1a3a6b")

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")

    for row, emp in enumerate(employees, 2):
        ws.cell(row=row, column=1, value=emp.employee_name)
        ws.cell(row=row, column=2, value=emp.designation or "")
        ws.cell(row=row, column=3,
                value=emp.organization.organization_name if emp.organization else "")
        ws.cell(row=row, column=4,
                value=emp.department.department_name if emp.department else "")
        ws.cell(row=row, column=5, value=emp.office_phone or "")
        ws.cell(row=row, column=6, value=emp.mobile_phone or "")
        ws.cell(row=row, column=7, value=emp.email or "")
        ws.cell(row=row, column=8, value="Yes" if emp.is_utility_head else "")

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 22

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="employees.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _export_employees_docx(employees, title="Employee Directory"):
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    heading = doc.add_heading(title, level=1)
    heading.runs[0].font.color.rgb = RGBColor(0x1a, 0x3a, 0x6b)

    table = doc.add_table(rows=1, cols=7)
    table.style = "Light List Accent 1"
    hdr = table.rows[0].cells
    for i, h in enumerate(
        ["Name", "Designation", "Organization", "Office Phone", "Mobile", "Email", "Utility Head"]
    ):
        hdr[i].text = h
        hdr[i].paragraphs[0].runs[0].bold = True

    for emp in employees:
        row = table.add_row().cells
        row[0].text = emp.employee_name or ""
        row[1].text = emp.designation or ""
        row[2].text = emp.organization.organization_name if emp.organization else ""
        row[3].text = emp.office_phone or ""
        row[4].text = emp.mobile_phone or ""
        row[5].text = emp.email or ""
        row[6].text = "Yes" if emp.is_utility_head else ""

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="employees.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


def _export_org_docx(org_name, address, employees, control_rooms):
    from docx import Document
    from docx.shared import RGBColor

    doc = Document()
    heading = doc.add_heading(org_name, level=1)
    heading.runs[0].font.color.rgb = RGBColor(0x1a, 0x3a, 0x6b)
    if address:
        doc.add_paragraph(address)

    if employees:
        doc.add_heading("Employees", level=2)
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Light List Accent 1"
        hdr = tbl.rows[0].cells
        for i, h in enumerate(["Name", "Designation", "Office Phone", "Mobile", "Email"]):
            hdr[i].text = h
            hdr[i].paragraphs[0].runs[0].bold = True
        for emp in employees:
            r = tbl.add_row().cells
            r[0].text = emp.employee_name or ""
            r[1].text = emp.designation or ""
            r[2].text = emp.office_phone or ""
            r[3].text = emp.mobile_phone or ""
            r[4].text = emp.email or ""

    if control_rooms:
        doc.add_heading("Control Room", level=2)
        tbl2 = doc.add_table(rows=1, cols=3)
        tbl2.style = "Light List Accent 1"
        hdr2 = tbl2.rows[0].cells
        for i, h in enumerate(["Name", "Phone Number", "Email"]):
            hdr2[i].text = h
            hdr2[i].paragraphs[0].runs[0].bold = True
        for cr in control_rooms:
            r = tbl2.add_row().cells
            r[0].text = cr.name or ""
            r[1].text = cr.phone_number or ""
            r[2].text = cr.email or ""

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in org_name)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{safe_name}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@user_bp.route("/export/utility-heads")
def export_utility_heads():
    fmt = request.args.get("format", "xlsx")
    heads = (
        Employee.query
        .filter_by(is_utility_head=True)
        .join(Organization)
        .order_by(Organization.organization_name, Employee.employee_name)
        .all()
    )
    if fmt == "docx":
        return _export_utility_heads_docx(heads)
    return _export_utility_heads_xlsx(heads)


def _export_utility_heads_xlsx(heads):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Utility Heads"

    headers = ["#", "Name", "Designation", "Organization", "Address",
               "Office Phone", "Mobile", "Email"]
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1a3a6b")

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")

    for row, emp in enumerate(heads, 2):
        ws.cell(row=row, column=1, value=row - 1)
        ws.cell(row=row, column=2, value=emp.employee_name or "")
        ws.cell(row=row, column=3, value=emp.designation or "")
        ws.cell(row=row, column=4,
                value=emp.organization.organization_name if emp.organization else "")
        ws.cell(row=row, column=5,
                value=(emp.organization.address or "") if emp.organization else "")
        ws.cell(row=row, column=6, value=emp.office_phone or "")
        ws.cell(row=row, column=7, value=emp.mobile_phone or "")
        ws.cell(row=row, column=8, value=emp.email or "")

    col_widths = [5, 28, 30, 45, 55, 22, 18, 30]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(i)
        ].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True,
        download_name="utility_heads.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _export_utility_heads_docx(heads):
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    heading = doc.add_heading("Utility Heads – Western Region", level=1)
    heading.runs[0].font.color.rgb = RGBColor(0x1a, 0x3a, 0x6b)

    table = doc.add_table(rows=1, cols=7)
    table.style = "Light List Accent 1"
    hdr = table.rows[0].cells
    for i, h in enumerate(
        ["Name", "Designation", "Organization", "Address",
         "Office Phone", "Mobile", "Email"]
    ):
        hdr[i].text = h
        hdr[i].paragraphs[0].runs[0].bold = True

    for emp in heads:
        row = table.add_row().cells
        row[0].text = emp.employee_name or ""
        row[1].text = emp.designation or ""
        row[2].text = emp.organization.organization_name if emp.organization else ""
        row[3].text = (emp.organization.address or "") if emp.organization else ""
        row[4].text = emp.office_phone or ""
        row[5].text = emp.mobile_phone or ""
        row[6].text = emp.email or ""

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True,
        download_name="utility_heads.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def _abort404():
    from flask import abort
    abort(404)
